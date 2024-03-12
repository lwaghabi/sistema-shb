VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTreinamentoConsulta 
   Caption         =   "frmAsoConsulta"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   6135
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
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
         TabIndex        =   0
         Top             =   480
         Width           =   5655
      End
      Begin VB.OptionButton optSeleciona 
         Caption         =   "Com data dentro do prazo de aviso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   5655
      End
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
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3000
      Width           =   8775
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
      Left            =   17280
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
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
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdConsultar 
      BackColor       =   &H0000FF00&
      Caption         =   "Consultar"
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
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid grdTreinamento 
      Height          =   4935
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   20295
      _ExtentX        =   35798
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   105
      FixedCols       =   0
      FormatString    =   $"frmTreinamentoConsulta.frx":0000
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
   Begin VB.Label Label1 
      Caption         =   "AGENDA DE PROGRAMAÇÃO DE CURSO/TREINAMENTO POR FUNCIONÁRIO "
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
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   15015
   End
   Begin VB.Label Label2 
      Caption         =   "Funcionário"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label3 
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
      Left            =   17400
      TabIndex        =   8
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "<<<<<-------------------Percentual Restante------------------>>>>>"
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
      Left            =   10560
      TabIndex        =   7
      Top             =   3720
      Width           =   9495
   End
End
Attribute VB_Name = "frmTreinamentoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataInvertida As String
Dim DataHojeInvertida As String
Dim DataAnterior As Date
Dim DataBase As Date
Dim DataDias As Date

Dim QtdDiasTotal As Double
Dim QtdDiasHoje As Double
Dim QtdDias As Double
Dim PercentDecorrido As Double

Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer

Dim AnoDb As Integer
Dim MesDb As Integer
Dim DiaDb As Integer

Dim ColaboradorAnterior As String

Dim Linha As Integer
Dim Ind As Integer
Dim IndCol As Integer
Dim IndCmb As Integer

Dim pessoaCmb(50) As String
Dim Encontrei As Integer



Private Sub cmdConsultar_Click()

Call LimpaGrdCurso

Call Rotina_AbrirBanco

agcto.Open "Select * from treinamentoagenda where chPessoa = ('" & cmbPessoa & "') and agctoDataProxCurso > ('" & DataHojeInvertida & "')", db, 3, 3
If agcto.EOF Then
   MsgBox ("Funcionário sem agenda de Curso/Treinamento."), vbInformation
   Call FechaDB
   Exit Sub
End If

Linha = 1

agcto.MoveFirst

Do While Not agcto.EOF
   grdTreinamento.Rows = Linha + 1
   If agcto!agctoDataProxCurso < Date Then
      QtdDias = 0
      Ind = 0
   Else
      QtdDiasTotal = agcto!agctoDataProxCurso - agcto!chDataTreinamento
      QtdDiasHoje = Date - agcto!chDataTreinamento
      
      QtdDias = agcto!agctoDataProxCurso - Date
      
      'x=(parc X 100) / tot
      
      PercentDecorrido = (QtdDiasHoje * 100) / QtdDiasTotal
      Ind = 100 - PercentDecorrido
   End If
   grdTreinamento.TextMatrix(Linha, 0) = agcto!chNomeCurso
   grdTreinamento.TextMatrix(Linha, 1) = agcto!chDataTreinamento
   grdTreinamento.TextMatrix(Linha, 2) = agcto!agctoDataProxCurso
   grdTreinamento.TextMatrix(Linha, 3) = QtdDias
   grdTreinamento.TextMatrix(Linha, 4) = Ind & "%"
   If QtdDias < 21 Then
      grdTreinamento.Col = 3
      grdTreinamento.Row = Linha
      grdTreinamento.CellBackColor = vbYellow
   Else
      grdTreinamento.Row = Linha
      grdTreinamento.Col = 3
      grdTreinamento.CellBackColor = Empty
   End If
   If Ind > 0 Then
      For IndCol = 1 To Ind
          grdTreinamento.Col = IndCol + 4
          grdTreinamento.Row = Linha
          
          If QtdDias < 21 Then
             grdTreinamento.CellBackColor = vbRed
          Else
             grdTreinamento.CellBackColor = vbGreen
          End If
                
      Next
   End If
   agcto.MoveNext
   Linha = Linha + 1
   
Loop

Call FechaDB

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtHoje = Date

optTodos = False
optSeleciona = False

ano = Year(Date)
mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

End Sub

Private Sub optSeleciona_Click()
txtHoje = Date
ColaboradorAnterior = Empty
DataAnterior = Empty

ano = Year(Date)
mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

cmbPessoa.Clear

Call Rotina_AbrirBanco

Call LimpaGrdCurso

agcto.Open "Select * from treinamentoagenda where agctoStatus = ('" & 0 & "')", db, 3, 3
If agcto.EOF Then
   Call FechaDB
   Exit Sub
End If

Linha = 1

agcto.MoveFirst

Do While Not agcto.EOF
   If cto.State = 1 Then
      cto.Close: Set cto = Nothing
   End If
   
   cto.Open "Select * from treinamento where chNomeCurso = ('" & agcto!chNomeCurso & "')", db, 3, 3
   If Not cto.EOF Then
      DataDias = Date + cto!ctoAvisoEm
      ano = Year(DataDias)
      mes = Month(DataDias)
      Dia = Day(DataDias)
      DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
   Else
      MsgBox ("Curso/Treinamento não encontrado. Comunicar ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   AnoDb = Year(agcto!agctoDataProxCurso)
   MesDb = Month(agcto!agctoDataProxCurso)
   DiaDb = Day(agcto!agctoDataProxCurso)

   DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")

   If DataInvertida > DataHojeInvertida Then
      If Not (DataInvertida > DataBase) Then
         If Not (agcto!chPessoa = ColaboradorAnterior) Then
            cmbPessoa.AddItem agcto!chPessoa
            ColaboradorAnterior = agcto!chPessoa
            DataAnterior = agcto!agctoDataProxCurso
         End If
      End If
   End If
   agcto.MoveNext

Loop

End Sub

Private Sub optTodos_Click()

Call Rotina_AbrirBanco

cmbPessoa.Clear

pes.Open "Select * from pessoa where pesRazaoSocial > ('" & Empty & "') and pesTipoPessoa = ('" & 6 & "') and pesStatusPessoa = ('" & 0 & "')", db, 3, 3
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
Call FechaDB

End Sub

Public Sub LimpaGrdCurso()
   grdTreinamento.Rows = 1
End Sub

