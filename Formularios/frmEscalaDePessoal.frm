VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEscalaDePessoal 
   Caption         =   "frmEscalaDePessoal"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtMesRef 
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
      Height          =   660
      Left            =   9750
      TabIndex        =   25
      Top             =   3000
      Width           =   9855
   End
   Begin VB.TextBox txtUnidadeOperacional 
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox txtQtdDias 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7320
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir 
      BackColor       =   &H000000FF&
      Caption         =   "Excluir"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      BackColor       =   &H0000FF00&
      Caption         =   "Salvar"
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdLogistica 
      Height          =   6135
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   20295
      _ExtentX        =   35798
      _ExtentY        =   10821
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtReferencia 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "mm/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "mm/yyyy"
      Format          =   224198657
      CurrentDate     =   44651
      MaxDate         =   2958435
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   18000
      TabIndex        =   18
      Top             =   600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtFinalEvento 
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   224198657
      CurrentDate     =   44648
   End
   Begin MSComCtl2.DTPicker dtInicioEvento 
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   224198657
      CurrentDate     =   44648
   End
   Begin VB.ComboBox cmbEvento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.ComboBox cmbColaborador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.ComboBox cmbUnidadeOperacional 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.ComboBox cmbPessoa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label12 
      Caption         =   "Referência"
      Height          =   375
      Left            =   9750
      TabIndex        =   26
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   "Unidade Operacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "Qtd Dias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Referência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      Height          =   375
      Left            =   18000
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Final do Evento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Início do Evento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Evento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Colaborador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   13
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Unidade Operacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "CONTROLE DE LOGÍSTICA DE PESSOAL"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmEscalaDePessoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim NumDiasMes As Integer
Dim DataInicioInvertida As Date
Dim DataFinalInvertida As Date
Dim DataInvertida As String
Dim DataHoje As Date
Dim MesProximo As Integer
Dim IndLinha As Integer
Dim IndCol As Integer
Dim Ind As Integer
Dim CabecalhoDias As String
Dim AnoMesRef As String
Dim ColaboradorAnterior As String
Dim CodEvento As String
Dim Resp As String
Dim Dia As Integer
Dim Mes As Integer
Dim ano As Integer
Dim DiasProxMes As Integer
Dim DataFimProxMes As String
Dim InicioDataChLeitura As String
Dim FinalDataChLeitura As String
Dim DataProximoMes As Date



Private Sub cmbColaborador_LostFocus()

If cmbColaborador = " TODOS" Then
   Call CarregaGrid
End If

End Sub

Private Sub cmbPessoa_LostFocus()

Call Rotina_AbrirBanco

cmbUnidadeOperacional.Clear

uoper.Open "Select * from UnidadeOperacional where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If uoper.EOF Then
   MsgBox ("Cliente sem Unidade Operacional cadastrada."), vbInformation
   Call FechaDB
   Exit Sub
End If

uoper.MoveFirst

Do While Not uoper.EOF

   cmbUnidadeOperacional.AddItem uoper!chUnidadeOperacional

   uoper.MoveNext
   
Loop

Call FechaDB

End Sub

Private Sub cmbUnidadeOperacional_LostFocus()

Call Rotina_AbrirBanco

cmbColaborador.Clear

pes.Open "Select * from Pessoa where pesClienteLocador = ('" & cmbPessoa & "') and pesUnidadeOperacional = ('" & cmbUnidadeOperacional & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Não há Colaborador Cadastrado para este Cliente/Unidade Operacional."), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

cmbColaborador.AddItem " TODOS"

Do While Not pes.EOF
   
   cmbColaborador.AddItem pes!chPessoa
   
   pes.MoveNext
   
Loop
   
cmbColaborador.ListIndex = 0

txtUnidadeOperacional = cmbUnidadeOperacional

Call FechaDB

End Sub

Private Sub cmdExcluir_Click()

AnoMesRef = Year(dtReferencia) & Format$(Month(dtReferencia), "00")

Call Rotina_AbrirBanco

'DataInvertida = Year(dtInicioEvento)

ano = Year(dtInicioEvento)
Mes = Format$(Month(dtInicioEvento), "00")
Dia = Format$(Day(dtInicioEvento), "00")

DataInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "') and chColaborador = ('" & cmbColaborador & "') and chEvento = ('" & cmbEvento & "') and chInicioEvento = ('" & DataInvertida & "')", db, 3, 3

If lgt.EOF Then
   MsgBox ("Exclusão de Programação inexistente."), vbCritical
   Call FechaDB
   Exit Sub
End If

DataProximoMes = lgt!lgtFimEventoReal

Resp = MsgBox("Exclusão de programação solicitada. Confirma???", vbExclamation + vbYesNo)

If Resp = vbYes Then
   lgt.Delete
   MsgBox ("Programação deletada com sucesso."), vbInformation
   Call CarregaGrid
End If

If Not Month(DataProximoMes) = Month(Date) Then
   If lgt.State = 1 Then
      lgt.Close: Set lgt = Nothing
   End If
   
   AnoMesRef = Year(DataProximoMes) & Format$(Month(DataProximoMes), "00")

   ano = Year(DataProximoMes)
   Mes = Format$(Month(DataProximoMes), "00")
   Dia = Format$("01", "00")
   
   DataInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
      
   Call Rotina_AbrirBanco
   
   lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "') and chColaborador = ('" & cmbColaborador & "') and chEvento = ('" & cmbEvento & "') and chInicioEvento = ('" & DataInvertida & "')", db, 3, 3
   
   If lgt.EOF Then
      MsgBox ("Exclusão de Programação de mes posterior inexistente."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   DataProximoMes = lgt!lgtFimEventoReal
   
   'Resp = MsgBox("Exclusão de programação solicitada. Confirma???", vbExclamation + vbYesNo)
   
   If Resp = vbYes Then
      lgt.Delete
      MsgBox ("Continuação da Programação no mes posterior deletada com sucesso."), vbInformation
      'Call CarregaGrid
   End If
End If

Call FechaDB

End Sub

Private Sub cmdSalvar_Click()

If cmbColaborador = " TODOS" Then
   MsgBox ("Somente salvar programação para um colaborador específico."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

AnoMesRef = Year(dtReferencia) & Format$(Month(dtReferencia), "00")

Call Rotina_AbrirBanco

ano = Year(dtInicioEvento)
Mes = Format$(Month(dtInicioEvento), "00")
Dia = Format$(Day(dtInicioEvento), "00")

DataInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "') and chColaborador = ('" & cmbColaborador & "') and chEvento = ('" & cmbEvento & "') and chInicioEvento = ('" & DataInvertida & "')", db, 3, 3

If lgt.EOF Then
   lgt.AddNew
End If

Ativ.Open "Select * FROM Atividade where atvAtividade = ('" & cmbEvento & "')", db, 3, 3
If Ativ.EOF Then
   MsgBox ("Evento não cadastrado. Cadastrar o evento e retornar para esta função."), vbInformation
   Call FechaDB
   Exit Sub
End If

lgt!chAnoMesRef = AnoMesRef
lgt!chPessoa = cmbPessoa
lgt!chUnidadeOperacional = cmbUnidadeOperacional
lgt!chColaborador = cmbColaborador
lgt!chEvento = cmbEvento
lgt!lgtCodEvento = Ativ!atvCodigoAtividade
lgt!chInicioEvento = dtInicioEvento
lgt!lgtFimEvento = dtFinalEvento
lgt!lgtFimEventoReal = DataFimProxMes
lgt!lgtStatusImport = 0

lgt.Update

Call FechaDB

Call CarregaGrid

If DiasProxMes > 0 Then
   Call GeraProximoMes
   DiasProxMes = 0
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dtReferencia_LostFocus()

If Month(dtReferencia) = 0 Then
   MsgBox ("Refrência informada inválida. Favor Informar o mês de refereência"), vbCritical
   dtReferencia.SetFocus
End If

ano = Year(dtReferencia)
Mes = Month(dtReferencia)
Dia = 1

AnoMesRef = ano & Format$(Mes, "00")

txtMesRef = Format$(dtReferencia, "mmm-yyyy")

Call LimpaGrid

Call GeraDataInicioDataFim

Call GeraGridlogistica

Call LimpaGrid

End Sub

Private Sub Form_Load()

dtInicioEvento = Date
dtFinalEvento = Date
txtHoje = Date
dtReferencia = Date

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where pesTipoPessoa = ('" & 0 & "')", db, 3, 3

pes.MoveFirst

Do While Not pes.EOF

   cmbPessoa.AddItem pes!chPessoa
   
   pes.MoveNext
   
Loop
 
Ativ.Open "Select * from Atividade", db, 3, 3

If Ativ.EOF Then
   MsgBox ("Tabela de Atividade vazia. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbEvento.Clear

Ativ.MoveFirst

Do While Not Ativ.EOF

   cmbEvento.AddItem Ativ!atvAtividade
   
   Ativ.MoveNext
   
Loop

Call FechaDB

End Sub

Public Sub GeraDataInicioDataFim()
Dim MesProximo As Integer

DataInicioInvertida = Format$(ano & "-" & Mes & "-" & Dia, "yyyy-mm-dd")

MesProximo = Format$(Mes, "00")
DataHoje = dtReferencia
Do While Mes = MesProximo
   DataHoje = DataHoje + 1
   MesProximo = Format$(Month(DataHoje), "00")
Loop

InicioDataChLeitura = DataInicioInvertida - 1

ano = Year(InicioDataChLeitura)
Mes = Month(InicioDataChLeitura)
Dia = Day(InicioDataChLeitura)

InicioDataChLeitura = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

ano = Year(DataHoje)
Mes = Month(DataHoje)
Dia = Day(DataHoje)

FinalDataChLeitura = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

DataHoje = DataHoje - 1

DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")

NumDiasMes = (DataFinalInvertida + 1) - DataInicioInvertida

End Sub

Public Sub GeraGridlogistica()

grdLogistica.Cols = NumDiasMes + 5

Ind = 1

CabecalhoDias = Empty

Do While Ind < NumDiasMes + 1
   CabecalhoDias = CabecalhoDias & ("|" & Format$(Ind, "00"))
   Ind = Ind + 1
Loop

grdLogistica.FormatString = "|" & "Colaborador                                         " & "|" & "Evento                                  " & "|" & "Inicio Evento" & "|" & "Final Evento" & CabecalhoDias

If NumDiasMes = 28 Then
   txtMesRef.Width = 9190
Else
   If NumDiasMes = 29 Then
      txtMesRef.Width = 99550
   Else
      If NumDiasMes = 30 Then
         txtMesRef.Width = 9940
      Else
         If NumDiasMes = 31 Then
            txtMesRef.Width = 10250
         End If
      End If
   End If
End If

End Sub

Public Sub CarregaGrid()

Call LimpaGrid

Call Rotina_AbrirBanco

lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "')", db, 3, 3
If lgt.EOF Then
   Call FechaDB
   Exit Sub
End If

lgt.MoveFirst

IndLinha = 1
IndCol = 0
ColaboradorAnterior = Empty

Do While Not lgt.EOF

   grdLogistica.Rows = IndLinha + 1
   If Not lgt!chColaborador = ColaboradorAnterior Then
      grdLogistica.TextMatrix(IndLinha, 1) = lgt!chColaborador
      ColaboradorAnterior = lgt!chColaborador
   Else
      grdLogistica.TextMatrix(IndLinha, 1) = Empty
   End If
   
   grdLogistica.TextMatrix(IndLinha, 0) = lgt!chColaborador
   grdLogistica.TextMatrix(IndLinha, 2) = lgt!chEvento
   grdLogistica.TextMatrix(IndLinha, 3) = lgt!chInicioEvento
   grdLogistica.TextMatrix(IndLinha, 4) = lgt!lgtFimEvento
   
   For IndCol = Day(lgt!chInicioEvento) To Day(lgt!lgtFimEvento)
       grdLogistica.TextMatrix(IndLinha, (IndCol + 4)) = lgt!lgtCodEvento
       grdLogistica.Col = (IndCol + 4)
       grdLogistica.Row = IndLinha

       If lgt!lgtCodEvento = "H" Then
          grdLogistica.CellBackColor = vbGreen
       Else
          If lgt!lgtCodEvento = "E" Or lgt!lgtCodEvento = "EM" Or lgt!lgtCodEvento = "EMB" Then
            grdLogistica.CellBackColor = vbRed
          Else
             If lgt!lgtCodEvento = "R" Or lgt!lgtCodEvento = "FLG" Or lgt!lgtCodEvento = "FOL" Or lgt!lgtCodEvento = "RR" Then
                grdLogistica.CellBackColor = vbBlue
             Else
                If lgt!lgtCodEvento = "F" Or lgt!lgtCodEvento = "FER" Then
                   grdLogistica.CellBackColor = vbYellow
                Else
                   grdLogistica.TextMatrix(IndLinha, (IndCol + 4)) = lgt!lgtCodEvento
                   grdLogistica.CellBackColor = &HFF00FF
                End If
             End If
           End If
        End If
        
       grdLogistica.CellForeColor = &H80000004
   Next
   
   IndLinha = IndLinha + 1
   
   lgt.MoveNext

Loop
 
End Sub


Private Sub grdlogistica_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

IndLinha = grdLogistica.Row
IndCol = grdLogistica.Col

If IndLinha > grdLogistica.RowSel Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

If grdLogistica.TextMatrix(IndLinha, 2) = Empty Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

cmbColaborador = grdLogistica.TextMatrix(IndLinha, 0)
cmbEvento = grdLogistica.TextMatrix(IndLinha, 2)
dtInicioEvento = grdLogistica.TextMatrix(IndLinha, 3)
dtFinalEvento = grdLogistica.TextMatrix(IndLinha, 4)
txtQtdDias = dtFinalEvento - (dtInicioEvento - 1)

txtQtdDias.SetFocus

End Sub


Private Sub txtQtdDias_LostFocus()

If txtQtdDias = Empty Then
   MsgBox ("Quantidade de dias não informado."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

dtFinalEvento = dtInicioEvento + (txtQtdDias - 1)
DataFimProxMes = dtFinalEvento

If Not (Month(dtFinalEvento) = Month(dtReferencia)) Then
   DiasProxMes = dtFinalEvento - DataFinalInvertida
   dtFinalEvento = DataHoje
End If

End Sub
Public Sub LimpaGrid()

   grdLogistica.Rows = 2

    grdLogistica.TextMatrix(1, 0) = Empty
    grdLogistica.TextMatrix(1, 1) = Empty
    
       If grdLogistica.Cols > 2 Then
       grdLogistica.TextMatrix(1, 2) = Empty
       grdLogistica.TextMatrix(1, 3) = Empty
       grdLogistica.TextMatrix(1, 4) = Empty
       For Ind = 5 To NumDiasMes + 4
           grdLogistica.TextMatrix(1, Ind) = Empty
           grdLogistica.Col = Ind
           grdLogistica.Row = 1
           grdLogistica.CellBackColor = Empty
       Next
    End If
End Sub


Public Sub GeraProximoMes()

AnoMesRef = Year(DataFimProxMes) & Format$(Month(DataFimProxMes), "00")

Call Rotina_AbrirBanco

ano = Year(DataFimProxMes)
Mes = Format$(Month(DataFimProxMes), "00")
Dia = Format$(1, "00")

DataInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "') and chColaborador = ('" & cmbColaborador & "') and chEvento = ('" & cmbEvento & "') and chInicioEvento = ('" & DataInvertida & "')", db, 3, 3

If lgt.EOF Then
   lgt.AddNew
End If

Ativ.Open "Select * FROM Atividade where atvAtividade = ('" & cmbEvento & "')", db, 3, 3
If Ativ.EOF Then
   MsgBox ("Evento não cadastrado. Cadastrar o evento e retornar para esta função."), vbInformation
   Call FechaDB
   Exit Sub
End If

lgt!chAnoMesRef = AnoMesRef
lgt!lgtTipo = 1
lgt!chPessoa = cmbPessoa
lgt!chUnidadeOperacional = cmbUnidadeOperacional
lgt!chColaborador = cmbColaborador
lgt!chEvento = cmbEvento
lgt!lgtCodEvento = Ativ!atvCodigoAtividade
lgt!chInicioEvento = DataInvertida
lgt!lgtFimEvento = DataFimProxMes
lgt!lgtFimEventoReal = DataFimProxMes
lgt!lgtStatusImport = 0

lgt.Update

Call FechaDB


End Sub
