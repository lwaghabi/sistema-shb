VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAsoConsulta 
   Caption         =   "frmAsoConsulta"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20205
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
   ScaleHeight     =   9675
   ScaleWidth      =   20205
   StartUpPosition =   2  'CenterScreen
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
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
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
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
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
      Left            =   17400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid grdAgenda 
      Height          =   5535
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   105
      FixedCols       =   0
      FormatString    =   $"frmAsoConsulta.frx":0000
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
   Begin VB.ComboBox cmbPessoa 
      Height          =   480
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Consulta"
      Height          =   1815
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   6135
      Begin VB.OptionButton optSeleciona 
         Caption         =   "Com data dentro do prazo de aviso"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   5655
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "<<<<<-------------------Percentual Restante------------------>>>>>"
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   3720
      Width           =   9015
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
      Left            =   17520
      TabIndex        =   9
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Funcionário"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "AGENDA DE PROGRAMAÇÃO DE EXAMES POR FUNCIONÁRIO - ASO"
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
      TabIndex        =   5
      Top             =   120
      Width           =   12495
   End
End
Attribute VB_Name = "frmAsoConsulta"
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

Call LimpaGrdAgenda

Call Rotina_AbrirBanco

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and asoaStatus = ('" & 0 & "')", db, 3, 3
If asoa.EOF Then
   MsgBox ("Funcionário sem agenda de exames."), vbInformation
   Call FechaDB
   Exit Sub
End If

Linha = 1

asoa.MoveFirst

Do While Not asoa.EOF
   grdAgenda.Rows = Linha + 1
   QtdDiasTotal = asoa!asoaDataProxExame - asoa!chDataExame
   QtdDiasHoje = Date - asoa!chDataExame
   QtdDias = asoa!asoaDataProxExame - Date
   'x=(parc X 100) / tot
   
   PercentDecorrido = (QtdDiasHoje * 100) / QtdDiasTotal
   Ind = 100 - PercentDecorrido
   If Ind < 1 Then
      Ind = 0
   End If
   grdAgenda.TextMatrix(Linha, 0) = asoa!chNomeExame
   grdAgenda.TextMatrix(Linha, 1) = asoa!chDataExame
   grdAgenda.TextMatrix(Linha, 2) = asoa!asoaDataProxExame
   grdAgenda.TextMatrix(Linha, 3) = QtdDias
   grdAgenda.TextMatrix(Linha, 4) = Ind & "%"
   If QtdDias < 21 Then
      grdAgenda.Col = 3
      grdAgenda.Row = Linha
      grdAgenda.CellBackColor = vbYellow
   Else
      grdAgenda.Row = Linha
      grdAgenda.Col = 3
      grdAgenda.CellBackColor = Empty
   End If
   For IndCol = 1 To Ind
   
       grdAgenda.Col = IndCol + 4
       grdAgenda.Row = Linha
       
       If QtdDias < 21 Then
          grdAgenda.CellBackColor = vbRed
       Else
          grdAgenda.CellBackColor = vbGreen
       End If
            
  Next
   asoa.MoveNext
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

Call LimpaGrdAgenda

asoa.Open "Select * from asoagenda where asoaStatus = ('" & 0 & "')", db, 3, 3
If asoa.EOF Then
   Call FechaDB
   Exit Sub
End If

Linha = 1

asoa.MoveFirst

Do While Not asoa.EOF
   If asoe.State = 1 Then
      asoe.Close: Set asoe = Nothing
   End If
   
   asoe.Open "Select * from asoexame where chNomeExame = ('" & asoa!chNomeExame & "')", db, 3, 3
   If Not asoe.EOF Then
      If asoe!exmUnidTempo = 0 Then
         DataDias = Date + asoe!exmPrazoAviso
         ano = Year(DataDias)
         mes = Month(DataDias)
         Dia = Day(DataDias)
         DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
      Else
         If asoe!exmUnidTempo = 1 Then
            ano = Year(Date)
            mes = Month(Date)
            mes = mes + asoe!exmPrazoAviso
            If mes > 12 Then
               ano = Year(Date)
               ano = ano + 1
               mes = mes - 12
            End If
            Dia = Day(Date)
            DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
         Else
            ano = Year(Date)
            ano = ano + asoe!exmPrazoAviso
            mes = Month(Date)
            Dia = Day(Date)
            DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
         End If
      End If
   End If
   
   AnoDb = Year(asoa!asoaDataProxExame)
   MesDb = Month(asoa!asoaDataProxExame)
   DiaDb = Day(asoa!asoaDataProxExame)

   DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")

   'If DataInvertida > DataHojeInvertida Then
      If Not (DataInvertida > DataBase) Then
         If Not (asoa!chPessoa = ColaboradorAnterior) Then
            cmbPessoa.AddItem asoa!chPessoa
            ColaboradorAnterior = asoa!chPessoa
            DataAnterior = asoa!asoaDataProxExame
         End If
      End If
   'End If
   asoa.MoveNext

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

Public Sub LimpaGrdAgenda()
   grdAgenda.Rows = 1
End Sub
